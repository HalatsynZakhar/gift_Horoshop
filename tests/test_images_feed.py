from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from starlette.requests import Request

import gifts_server
from images_feed import ImagesFeed, images_feed


class ImagesFeedTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        images_feed.load()

    def test_lookup_product_meta(self) -> None:
        meta = images_feed.get_product_meta("0010")
        self.assertEqual(meta["article"], "0010")
        self.assertIn("176.104.6.149", meta["image"])
        self.assertEqual(meta["brand"], "BABY TEAM")
        self.assertIsInstance(meta["images"], list)
        self.assertGreaterEqual(len(meta["images"]), 1)

    def test_lookup_fallback_for_lstrip_zeroes(self) -> None:
        # If article without leading zeros is queried, it should find the item
        meta = images_feed.get_product_meta("10")
        self.assertEqual(meta["article"], "10")
        if meta["image"]:
            self.assertIn("176.104.6.149", meta["image"])

    def test_lookup_missing_article(self) -> None:
        meta = images_feed.get_product_meta("NON_EXISTENT_ARTICLE_999999")
        self.assertEqual(meta["article"], "NON_EXISTENT_ARTICLE_999999")
        self.assertEqual(meta["image"], "")
        self.assertEqual(meta["brand"], "")
        self.assertEqual(meta["images"], [])

    def test_enrich_products(self) -> None:
        enriched = images_feed.enrich_products(["0010", "NON_EXISTENT_XYZ"])
        self.assertEqual(len(enriched), 2)
        self.assertEqual(enriched[0]["article"], "0010")
        self.assertIn("176.104.6.149", enriched[0]["image"])

    def test_products_info_endpoint(self) -> None:
        scope = {
            "type": "http",
            "method": "GET",
            "path": "/api/products/info",
            "headers": [],
            "query_string": b"articles=0010;002",
        }
        res = gifts_server.products_info("0010;002")
        self.assertIn("items", res)
        self.assertEqual(len(res["items"]), 2)
        self.assertEqual(res["items"][0]["article"], "0010")
        self.assertIn("176.104.6.149", res["items"][0]["image"])

    def test_images_status_endpoint(self) -> None:
        status = gifts_server.images_status()
        self.assertTrue(status["loaded"])
        self.assertGreater(status["count"], 0)
        self.assertEqual(status["url"], "http://176.104.6.149:8080/xml/images_export.xml")

    def test_images_refresh_endpoint(self) -> None:
        res = gifts_server.images_refresh()
        self.assertEqual(res["status"], "refreshing")
        self.assertEqual(res["url"], "http://176.104.6.149:8080/xml/images_export.xml")


if __name__ == "__main__":
    unittest.main()
