from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


def _colnum_to_letters(n: int) -> str:
    """Convert 1-based column number to Excel letters. 1->A, 27->AA."""
    if n < 1:
        raise ValueError("column must be >= 1")
    letters = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


class convert_row_col_to_a1(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        """Parameters:
        - row: int (1-based)
        - column: int (1-based)

        Returns JSON {"cell": "A1", "row": 1, "column": 1} and a text message with the A1 string.
        """
        row = tool_parameters.get("row", 1)
        column = tool_parameters.get("column", 1)

        try:
            if row is None or column is None:
                raise Exception("缺少必填参数：row 和 column")

            row_i = int(row)
            col_i = int(column)

            if row_i < 1 or col_i < 1:
                raise Exception("row 和 column 必须为大于等于 1 的整数")

            cell_name = f"{_colnum_to_letters(col_i)}{row_i}"

            yield self.create_json_message({"cell": cell_name, "row": row_i, "column": col_i})
            yield self.create_text_message(cell_name)
        except Exception as e:
            raise Exception(f"行列转单元格失败：{e}")
