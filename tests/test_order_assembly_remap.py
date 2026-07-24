import json
import unittest

import api_server


def item(product_id, article, name, quantity=1, unit="piece", buyer_order=""):
    return (product_id, article, name, quantity, unit, buyer_order, None)


class OrderAssemblyRemapTests(unittest.TestCase):
    def test_deleting_middle_line_keeps_survivors_on_original_pallets(self):
        old_items = [
            item(1, "A", "Product A"),
            item(2, "B", "Product B"),
            item(3, "C", "Product C"),
        ]
        new_items = [old_items[0], old_items[2]]
        state = {
            "pallets": [
                {
                    "id": 1,
                    "palletNumber": "1",
                    "slots": [
                        {"lineIndex": 0, "directQty": 10},
                        {"lineIndex": 1, "directQty": 20},
                    ],
                },
                {
                    "id": 2,
                    "palletNumber": "2",
                    "slots": [{"lineIndex": 2, "directQty": 30}],
                },
            ]
        }

        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state), old_items, new_items
            )
        )

        self.assertEqual(
            result["pallets"][0]["slots"],
            [{"lineIndex": 0, "directQty": 10}],
        )
        self.assertEqual(
            result["pallets"][1]["slots"],
            [{"lineIndex": 1, "directQty": 30}],
        )

    def test_quantity_edit_does_not_detach_slot(self):
        old_items = [item(1, "A", "Product A", quantity=10)]
        new_items = [item(1, "A", "Product A", quantity=25)]
        state = {
            "pallets": [
                {"id": 1, "slots": [{"lineIndex": 0, "directQty": 5, "batchNumber": "BN"}]}
            ]
        }

        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state), old_items, new_items
            )
        )

        self.assertEqual(result["pallets"][0]["slots"][0]["lineIndex"], 0)
        self.assertEqual(result["pallets"][0]["slots"][0]["directQty"], 5)
        self.assertEqual(result["pallets"][0]["slots"][0]["batchNumber"], "BN")

    def test_unit_and_quantity_edit_keeps_slots(self):
        """Ригла и др.: смена qty + штуки→наборы не должна очищать товары на паллетах."""
        old_items = [item(1, "A", "Product A", quantity=10, unit="piece")]
        new_items = [item(1, "A", "Product A", quantity=25, unit="set")]
        state = {
            "pallets": [
                {
                    "id": 1,
                    "palletNumber": "1",
                    "slots": [
                        {
                            "lineIndex": 0,
                            "directQty": "3",
                            "batchNumber": "KEEP",
                            "mode": "direct",
                            "directUnit": "box",
                        }
                    ],
                },
                {
                    "id": 2,
                    "palletNumber": "2",
                    "slots": [
                        {
                            "lineIndex": 0,
                            "directQty": "1",
                            "batchNumber": "KEEP2",
                        }
                    ],
                },
            ]
        }
        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state), old_items, new_items
            )
        )
        self.assertEqual(len(result["pallets"]), 2)
        self.assertEqual(result["pallets"][0]["slots"][0]["batchNumber"], "KEEP")
        self.assertEqual(result["pallets"][1]["slots"][0]["batchNumber"], "KEEP2")
        self.assertEqual(result["pallets"][0]["palletNumber"], "1")

    def test_quantity_edit_preserves_slots_when_article_name_drift(self):
        """product_id совпадает — сборка не сбрасывается из‑за пустого article/name."""
        old_items = [item(1, "A", "Product A", quantity=10)]
        new_items = [item(1, "", "", quantity=25)]
        state = {
            "pallets": [
                {
                    "id": 1,
                    "slots": [
                        {
                            "lineIndex": 0,
                            "directQty": "3",
                            "batchNumber": "KEEP",
                        }
                    ],
                }
            ]
        }
        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state), old_items, new_items
            )
        )
        self.assertEqual(len(result["pallets"][0]["slots"]), 1)
        self.assertEqual(result["pallets"][0]["slots"][0]["batchNumber"], "KEEP")

    def test_merged_client_indices_follow_product_name_groups(self):
        old_items = [
            item(1, "A1", "Same product", buyer_order="1"),
            item(1, "A1", "Same product", buyer_order="2"),
            item(2, "B", "Other product", buyer_order="1"),
        ]
        new_items = [old_items[2]]
        state = {
            "assembly_line_indices": "merged",
            "pallets": [
                {
                    "id": 1,
                    "slots": [
                        {"lineIndex": 0, "directQty": 10},
                        {"lineIndex": 1, "directQty": 20},
                    ],
                }
            ],
        }

        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state),
                old_items,
                new_items,
                merge_by_name=True,
            )
        )

        self.assertEqual(
            result["pallets"][0]["slots"],
            [{"lineIndex": 0, "directQty": 20}],
        )

    def test_merged_qty_edit_keeps_all_slots(self):
        old_items = [
            item(1, "A1", "Same product", quantity=5, buyer_order="1"),
            item(1, "A1", "Same product", quantity=3, buyer_order="2"),
            item(2, "B", "Other product", quantity=10, buyer_order="1"),
        ]
        new_items = [
            item(1, "A1", "Same product", quantity=8, buyer_order="1"),
            item(1, "A1", "Same product", quantity=3, buyer_order="2"),
            item(2, "B", "Other product", quantity=10, buyer_order="1"),
        ]
        state = {
            "assembly_line_indices": "merged",
            "pallets": [
                {
                    "id": 1,
                    "slots": [
                        {"lineIndex": 0, "directQty": 10, "batchNumber": "X"},
                        {"lineIndex": 1, "directQty": 20, "batchNumber": "Y"},
                    ],
                }
            ],
        }
        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state),
                old_items,
                new_items,
                merge_by_name=True,
            )
        )
        self.assertEqual(result["pallets"][0]["slots"][0]["batchNumber"], "X")
        self.assertEqual(result["pallets"][0]["slots"][1]["batchNumber"], "Y")

    def test_merged_delete_one_raw_line_of_group_keeps_other_product_slot(self):
        """Удаление одной raw-строки группы не должно сбрасывать слот другого товара."""
        old_items = [
            item(1, "A1", "Same product", buyer_order="1"),
            item(1, "A1", "Same product", buyer_order="2"),
            item(2, "B", "Other product", buyer_order="1"),
        ]
        new_items = [
            item(1, "A1", "Same product", buyer_order="1"),
            item(2, "B", "Other product", buyer_order="1"),
        ]
        state = {
            "assembly_line_indices": "merged",
            "pallets": [
                {
                    "id": 1,
                    "slots": [
                        {"lineIndex": 0, "directQty": 10, "batchNumber": "X"},
                        {"lineIndex": 1, "directQty": 20, "batchNumber": "Y"},
                    ],
                }
            ],
        }
        result = json.loads(
            api_server._remap_assemble_state_after_order_item_edit(
                json.dumps(state),
                old_items,
                new_items,
                merge_by_name=True,
            )
        )
        self.assertEqual(
            result["pallets"][0]["slots"],
            [
                {"lineIndex": 0, "directQty": 10, "batchNumber": "X"},
                {"lineIndex": 1, "directQty": 20, "batchNumber": "Y"},
            ],
        )


if __name__ == "__main__":
    unittest.main()
